VERSION 5.00
Begin VB.Form FrmSel14DProPyme 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contribuyentes Art. 14 D LIR Pro Pyme"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   420
      Picture         =   "FrmSel14DProPyme.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   5
      Top             =   540
      Width           =   585
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Top             =   960
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Sel 
      Caption         =   "Seleccionar..."
      Height          =   315
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   4635
      Begin VB.OptionButton Op_14DNro3 
         Caption         =   "Art 14 D N° 3 Régimen Pro Pyme  General"
         Height          =   315
         Left            =   420
         TabIndex        =   0
         Top             =   480
         Width           =   3735
      End
      Begin VB.OptionButton Op_14DNro8 
         Caption         =   "Art. 14 D N° 8 Régimen Pro Pyme Transparente"
         Height          =   315
         Left            =   420
         TabIndex        =   1
         Top             =   960
         Width           =   3795
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ATENCIÓN: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   7
      Top             =   2700
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Antes de ingresar recuerde calcular Proporcionalidad de IVA CF"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   3060
      Width           =   4545
   End
End
Attribute VB_Name = "FrmSel14DProPyme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Sel_Click()
   Dim Frm As Form
   
   Set Frm = New FrmSelLib14ter
   
   If Op_14DNro3 <> 0 Then
      Call Frm.FView(Op_14DNro3.Caption)
      
   ElseIf Op_14DNro8 Then
      Call Frm.FView(Op_14DNro8.Caption)
         
   End If
   
   Set Frm = Nothing
   
End Sub

Private Sub Form_Load()
   
   If Not gEmpresa.ProPymeGeneral Then
      Op_14DNro3.Enabled = False
   End If
   
   If Not gEmpresa.ProPymeTransp Then
      Op_14DNro8.Enabled = False
   End If
   
   
End Sub

Private Sub Op_14DNro3_DblClick()
   Call PostClick(Bt_Sel)

End Sub

Private Sub Op_14DNro8_DblClick()
   Call PostClick(Bt_Sel)

End Sub

Private Sub Op_14TerA_DblClick()
   Call PostClick(Bt_Sel)

End Sub
