VERSION 5.00
Begin VB.Form FrmSelSeguimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Auditoría"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   -600
      TabIndex        =   0
      Top             =   -120
      Width           =   8175
      Begin VB.CommandButton Bt_Cancel 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   4680
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Bt_OK 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   4680
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "De"
         Height          =   2655
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   3615
         Begin VB.OptionButton Op_Seguimiento 
            Caption         =   "Documento"
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   3
            Top             =   1680
            Width           =   2415
         End
         Begin VB.OptionButton Op_Seguimiento 
            Caption         =   "Comprobante (Nuevo)"
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   2
            Top             =   1200
            Width           =   2175
         End
         Begin VB.OptionButton Op_Seguimiento 
            Caption         =   "Comprobante"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   1
            Top             =   720
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "FrmSelSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Op_SEGUAUD = 1
Const Op_SEGUCOMP = 2
Const Op_SEGUDOC = 3

Dim lIdx As Integer

Private Sub Bt_Cancel_Click()
Unload Me
End Sub

Private Sub Bt_OK_Click()

Dim Frm As FrmAuditoria
Dim FrmDoc As FrmSeguimientoDoc
Dim FrmComp As FrmSeguimientoComp


    Select Case lIdx
    
            Case Op_SEGUAUD
                
                Set Frm = New FrmAuditoria
                Call Frm.FView
                Set Frm = Nothing
                
           Case Op_SEGUCOMP
            
                Set FrmComp = New FrmSeguimientoComp
                FrmComp.Show vbModal
                Set FrmComp = Nothing
                
            Case Op_SEGUDOC
            
                Set FrmDoc = New FrmSeguimientoDoc
                FrmDoc.Show vbModal
                Set FrmDoc = Nothing
         
    End Select


End Sub

Private Sub Op_Seguimiento_Click(Index As Integer)
lIdx = Index
End Sub
