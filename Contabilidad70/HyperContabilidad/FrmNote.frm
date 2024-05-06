VERSION 5.00
Begin VB.Form FrmNote 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "X"
      Height          =   195
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   195
   End
   Begin VB.TextBox Tx_Note 
      BackColor       =   &H00C0FFFF&
      Height          =   1635
      Left            =   0
      MaxLength       =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "FrmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTxtNote As String
Dim lOper As Integer

Public Function FEdit(TxtNote As String)

   lOper = O_EDIT

   lTxtNote = TxtNote
   
   If lTxtNote = "" Then
      lTxtNote = gUsuario.Nombre & ":"
   End If
   
   Me.Show vbModal
   
   TxtNote = lTxtNote
   
End Function
Public Function FView(TxtNote As String)

   lOper = O_VIEW
   
   lTxtNote = TxtNote
      
   Me.Show vbModal
      
End Function

Private Sub bt_Cerrar_Click()

   lTxtNote = Tx_Note
   Unload Me
   
End Sub

Private Sub Form_Load()
   
   Tx_Note = lTxtNote & " "
   Tx_Note.SelStart = Len(Tx_Note) + 1
   If lOper = O_VIEW Then
      Tx_Note.Locked = True
   End If
   
End Sub

Private Sub Tx_Note_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyEscape Then
      Call bt_Cerrar_Click
   End If
   
End Sub
