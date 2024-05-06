VERSION 5.00
Begin VB.Form FrmCambioEstadoComp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Estado Comprobantes Seleccionados"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1380
      TabIndex        =   4
      Top             =   480
      Width           =   3855
      Begin VB.ComboBox Cb_Estado 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Estado:"
         Height          =   195
         Index           =   3
         Left            =   540
         TabIndex        =   5
         Top             =   480
         Width           =   1065
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton Bt_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   960
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   480
      Picture         =   "FrmCambioEstadoComp.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "FrmCambioEstadoComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lRc As Integer
Dim lEstado As Integer

Private Sub Bt_Cancel_Click()
    lRc = vbCancel
    Unload Me

End Sub

Private Sub Bt_OK_Click()
    lRc = vbOK
    lEstado = CbItemData(Cb_Estado)
    Unload Me
End Sub

Private Sub Form_Load()

    Call CbAddItem(Cb_Estado, gEstadoComp(EC_APROBADO), EC_APROBADO)
    Call CbAddItem(Cb_Estado, gEstadoComp(EC_PENDIENTE), EC_PENDIENTE)

    Cb_Estado.ListIndex = 0
    lEstado = 0
    
End Sub

Public Function FEdit(Estado As Integer) As Integer
    Me.Show vbModal
    Estado = lEstado
    FEdit = lRc
End Function
