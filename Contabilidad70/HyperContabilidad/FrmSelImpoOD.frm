VERSION 5.00
Begin VB.Form FrmSelImpoOD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importación de Otros Documentos"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   -840
      TabIndex        =   0
      Top             =   -600
      Width           =   10215
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   6960
         TabIndex        =   5
         Top             =   1320
         Width           =   1275
      End
      Begin VB.CommandButton Bt_Sel 
         Caption         =   "Seleccionar..."
         Default         =   -1  'True
         Height          =   315
         Left            =   6960
         TabIndex        =   4
         Top             =   960
         Width           =   1275
      End
      Begin VB.Frame Frame2 
         Caption         =   "Seleecione que desea Importar"
         Height          =   1695
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   5415
         Begin VB.OptionButton Op_OD 
            Caption         =   "Otros Documentos Full"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   3
            Top             =   1080
            Width           =   2895
         End
         Begin VB.OptionButton Op_OD 
            Caption         =   "Otros Documentos"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
      End
   End
End
Attribute VB_Name = "FrmSelImpoOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ODF_NORMAL = 0
Const ODF_FULL = 1

Private Sub Bt_Cancelar_Click()
Unload Me
End Sub

Private Sub Bt_Sel_Click()
Dim Frm As Form
    If Op_OD(ODF_NORMAL).Value <> 0 Then
       Set Frm = New FrmImpOtrosDocs
       Frm.Show vbModal
       Set Frm = Nothing
    Else
       Set Frm = New FrmImpOtrosDocFull
       Frm.Show vbModal
       Set Frm = Nothing
    
    End If
End Sub

Private Sub Form_Load()
Op_OD(ODF_NORMAL) = True
End Sub
