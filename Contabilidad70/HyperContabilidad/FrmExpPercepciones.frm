VERSION 5.00
Begin VB.Form FrmExpPercepciones 
   Caption         =   "Exportar Percepciones"
   ClientHeight    =   1410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   -600
      TabIndex        =   0
      Top             =   -840
      Width           =   7095
      Begin VB.CommandButton Bt_Exp 
         Caption         =   "Exportar"
         Height          =   315
         Left            =   4860
         TabIndex        =   6
         Top             =   1140
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Close 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   4860
         TabIndex        =   5
         Top             =   1500
         Width           =   1275
      End
      Begin VB.Frame Fr_Periodo 
         Caption         =   "Período"
         Height          =   975
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   2655
         Begin VB.TextBox Tx_Ano 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Año:"
            Height          =   195
            Index           =   1
            Left            =   660
            TabIndex        =   4
            Top             =   480
            Width           =   330
         End
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   840
         Picture         =   "FrmExpPercepciones.frx":0000
         ScaleHeight     =   570
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1140
         Width           =   585
      End
   End
End
Attribute VB_Name = "FrmExpPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bt_Close_Click()
Unload Me
End Sub

Private Sub Bt_Exp_Click()
If Export_Percepciones("PERC") = 0 Then
 Unload Me
End If
End Sub

Private Sub Form_Load()
Tx_Ano = gEmpresa.Ano
End Sub
